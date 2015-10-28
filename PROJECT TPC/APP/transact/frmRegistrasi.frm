VERSION 5.00
Begin VB.Form frmRegistrasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FINGER REGISTRATION"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "frmRegistrasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3405
      Index           =   1
      Left            =   143
      TabIndex        =   0
      Top             =   158
      Width           =   6075
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmRegistrasi.frx":000C
         Left            =   960
         List            =   "frmRegistrasi.frx":002E
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mulai"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   3
         Left            =   4500
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   795
         Width           =   1275
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   2
         Left            =   3060
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   780
         Width           =   1275
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   1
         Left            =   1620
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   795
         Width           =   1275
      End
      Begin VB.PictureBox picSample 
         Height          =   1635
         Index           =   0
         Left            =   180
         ScaleHeight     =   1575
         ScaleWidth      =   1215
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   780
         Width           =   1275
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   2535
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   2535
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   2100
         Shape           =   3  'Circle
         Top             =   2535
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape dot 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   660
         Shape           =   3  'Circle
         Top             =   2535
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Jari"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   307
         Width           =   585
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   3000
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmRegistrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Registration As UareUSDK.clsFPRegistration
Attribute Registration.VB_VarHelpID = -1
Dim WithEvents FPImage As UareUSDK.clsFPImage
Attribute FPImage.VB_VarHelpID = -1
Dim Finger As String




Private Sub Command1_Click()
  If frmKaryawan.ID.Text = "" Then
     MsgBox "ID Harus Di isi ", -1, "Perhatian"
     Exit Sub
  End If
  
  Command1.Enabled = False
  Registration.FPRegistration frmKaryawan.ID.Text, GetFingerNumber(Combo2.Text)
  Label10.Caption = " Silahkan Letakkan jari anda pada Sensor "
End Sub

Private Sub Exit_Click()
  Close
End Sub

Private Sub Registration_FPRegistrationImage(CurentSample As Integer)
  picSample(CurentSample) = LoadPicture(App.Path & "\FPTemp.BMP")
  dot(CurentSample).Visible = True
End Sub

Sub Registrasi()
  Set Registration = New UareUSDK.clsFPRegistration
  
  
  Registration.PictureSamplePath = App.Path & "\FPTempTTT.BMP"
  Registration.PictureSampleHeight = picSample(0).Height
  Registration.PictureSampleWidth = picSample(0).Width
  Command1.Enabled = True
  Combo2.Enabled = True
  If frmKaryawan.ID.Text = "" Then
  MsgBox "ID tidak Boleh Kosong", vbInformation, "Keterangan"
  End If
End Sub

Private Sub Registration_FPRegistrationStatus(Status As RegistrationStatus)
  Command1.Enabled = True
  Select Case Status
    Case r_OK
      MsgBox "Registration Success", 0, "Keterangan"
      
      Combo2.Text = "Left Pinkie"
      Command1.Enabled = False
      Combo2.Enabled = False
      Label10.Caption = ""
      
      picSample(0) = Nothing
      picSample(1) = Nothing
      picSample(2) = Nothing
      picSample(3) = Nothing
      
      dot(0).Visible = False
      dot(1).Visible = False
      dot(2).Visible = False
      dot(3).Visible = False
      
    Case r_FpIdAlreadyExist
      MsgBox "ID dan finger number already exist", 0, "Keterangan"
    Case r_NoDevice
      MsgBox "Device not exits", 0, "Keterangan"
    Case r_WrongDeviceSN
      MsgBox "Wrong Device Serial Number", 0, "Keterangan"
    Case r_RegFailed
      MsgBox "Registration Gagal"
    Case r_WrongFingerNr
      MsgBox "Finger number must between 0 to 9", 0, "Keterangan"
  End Select
End Sub

Private Function GetFingerNumber(Finger As String) As FingerNumber
  Dim j As Byte
  Select Case Finger
    Case "Jari Kelingking Kiri"
      j = 0
    Case "Jari Manis Kiri"
      j = 1
    Case "Jari Tengah Kiri"
      j = 2
    Case "Jari Telunjuk Kiri"
      j = 3
    Case "Ibu Jari Kiri"
      j = 4
    Case "Ibu Jari Kanan"
      j = 5
    Case "Jari Telunjuk Kanan"
      j = 6
    Case "Jari Tengah Kanan"
      j = 7
    Case "Jari Manis Kanan"
      j = 8
    Case "Jari Kelingking Kanan"
      j = 9
    Case "(None)"
      j = 10
    Case Else
      j = 10
  End Select
  
  GetFingerNumber = j
End Function

Private Sub Form_Load()
  Registrasi
End Sub
