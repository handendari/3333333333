VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Begin VB.Form frm_mst_device 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER DEVICE"
   ClientHeight    =   4035
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7530
   Icon            =   "frm_mst_device.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   375
      Left            =   720
      OleObjectBlob   =   "frm_mst_device.frx":058A
      TabIndex        =   6
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmd_test 
      Caption         =   "&Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_mst_device.frx":05AE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4200
      Picture         =   "frm_mst_device.frx":0B38
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   5760
      Picture         =   "frm_mst_device.frx":10C2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "FINGERPRINT DEVICE"
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   6135
      Begin VB.TextBox txt_ip_address 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP ADDRESS"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_mst_device"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cn_fp As Boolean
'Dim vEMachineNumber As Integer


Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()
With rs
    .Fields("ip_address").Value = Trim(txt_ip_address)
    .Update
End With

FG_IP_ADDRESS = Trim(txt_ip_address)
End Sub

Private Sub cmd_test_Click()
Call test_device_connection
End Sub

Private Sub Form_Load()

cn_fp = False
vMachineNumber = 1

Call load_data

Call load_data_user_access(Me)
cmd_save.Enabled = blnUser_Edit
End Sub

Private Sub load_data()
If rs.State = 1 Then rs.Close
rs.Open "select * from m_device", CnG, adOpenKeyset, adLockOptimistic

If rs.RecordCount > 0 Then
    txt_ip_address = rs.Fields("ip_address").Value
Else
    txt_ip_address = ""
End If
End Sub

' --------------------------------------------

Private Function connect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If

cn_fp = CZKEM1.Connect_Net(FG_IP_ADDRESS, CLng(FG_PORT_NUMBER))

If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, False
    connect = True
Else
    connect = False
    Exit Function
End If
End Function

Private Function disconnect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If
End Function

Private Sub test_device_connection()
'On Error GoTo err_capture

If Not connect Then
    MsgBox "Error connecting to device...", vbCritical, headerMSG
    Exit Sub
End If

MsgBox "Connecting to device successfully tested!", vbInformation, headerMSG

Call disconnect

Exit Sub
err_capture:
MsgBox "Error connecting to device2!", vbCritical, headerMSG
Call disconnect
End Sub

